#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
convert_height_gui_asl.py (ver. GUI53 – 写真/動画排他選択版)

変更点
• 写真撮影と動画撮影のチェックボックスを同時に選択できないように排他制御を追加
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
    "北": 0.00,
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
        idx_elem = pm.find("wpml:index", NS)
        if idx_elem is None:
            continue
        idx = int(idx_elem.text)
        for action in pm.findall(".//wpml:action", NS):
            func_elem = action.find("wpml:actionActuatorFunc", NS)
            if func_elem is not None and func_elem.text == "orientedShoot":
                param = action.find("wpml:actionActuatorFuncParam", NS)
                if param is not None:
                    pitch_elem = param.find("wpml:gimbalPitchRotateAngle", NS)
                    yaw_elem = param.find("wpml:gimbalYawRotateAngle", NS)
                    head_elem = param.find("wpml:aircraftHeading", NS)
                    if pitch_elem is not None and yaw_elem is not None:
                        original[idx] = {
                            "pitch": float(pitch_elem.text),
                            "yaw": float(yaw_elem.text),
                            "heading": float(head_elem.text) if head_elem is not None else None
                        }
                break
    return original

# --- KMZ ユーティリティ -----------------------------------------------------
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
                if f.lower().endswith(".wpml"):
                    continue
                full = os.path.join(root, f)
                rel = os.path.relpath(full, out_root)
                zf.write(full, rel)
    if os.path.exists(out_kmz):
        os.remove(out_kmz)
    os.rename(tmp, out_kmz)
    return out_kmz

# --- 偏差計算 ---------------------------------------------------------------
def calculate_deviation(reference_coords, current_coords):
    ref_lng, ref_lat, ref_alt = reference_coords
    cur_lng, cur_lat, cur_alt = current_coords
    return cur_lng - ref_lng, cur_lat - ref_lat, cur_alt - ref_alt

def check_deviation_safety(deviation):
    dev_lng, dev_lat, dev_alt = deviation
    if (abs(dev_lng) > DEVIATION_THRESHOLD["lng"] or
        abs(dev_lat) > DEVIATION_THRESHOLD["lat"] or
        abs(dev_alt) > DEVIATION_THRESHOLD["alt"]):
        return False, (f"偏差が20mを超えています:\n"
                       f"経度偏差: {dev_lng:.8f}°\n"
                       f"緯度偏差: {dev_lat:.8f}°\n"
                       f"標高偏差: {dev_alt:.2f}m")
    return True, None

# --- KML 変換 ---------------------------------------------------------------
def convert_kml(tree, offset, do_photo, do_video, video_suffix,
                do_gimbal, gimbal_pitch_angle, gimbal_pitch_mode,
                yaw_fix, yaw_angle, yaw_mode, speed, sensor_modes, hover_time,
                coordinate_deviation=None, original_angles=None):
    # (同じ変換ロジック。元コードに忠実)

    # 座標偏差補正
    if coordinate_deviation:
        dev_lng, dev_lat, dev_alt = coordinate_deviation
        for pm in tree.findall(".//kml:Placemark", NS):
            coords_elem = pm.find(".//kml:coordinates", NS)
            if coords_elem is not None and coords_elem.text:
                coords = coords_elem.text.strip().split(',')
                if len(coords) >= 2:
                    try:
                        lng = float(coords[0]) + dev_lng
                        lat = float(coords[1]) + dev_lat
                        coords_elem.text = f"{lng},{lat}"
                    except ValueError:
                        pass

    # 高度補正＋EGM96
    for pm in tree.findall(".//kml:Placemark", NS):
        for tag in ("height", "ellipsoidHeight"):
            el = pm.find(f"wpml:{tag}", NS)
            if el is not None and el.text:
                try:
                    current_height = float(el.text) + offset
                    if coordinate_deviation:
                        current_height += coordinate_deviation[2]
                    el.text = str(current_height)
                except:
                    pass
    gh = tree.find(".//wpml:globalHeight", NS)
    if gh is not None and gh.text:
        try:
            current_global_height = float(gh.text) + offset
            if coordinate_deviation:
                current_global_height += coordinate_deviation[2]
            gh.text = str(current_global_height)
        except:
            pass
    for hm in tree.findall(".//wpml:heightMode", NS):
        hm.text = "EGM96"

    # 速度設定
    for tag, path in [
        ("globalTransitionalSpeed", ".//wpml:missionConfig"),
        ("autoFlightSpeed", ".//kml:Folder")
    ]:
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

    # ヨー固定設定（元角度維持でない場合のみグローバル設定）
    if yaw_fix and yaw_mode != "original" and yaw_angle is not None:
        for hp in tree.findall(".//wpml:globalWaypointHeadingParam", NS):
            m = hp.find("wpml:waypointHeadingMode", NS)
            a = hp.find("wpml:waypointHeadingAngle", NS)
            if m is not None:
                m.text = "fixed"
            if a is not None:
                a.text = str(int(yaw_angle))
        for pm in tree.findall(".//kml:Placemark", NS):
            hp = pm.find("wpml:waypointHeadingParam", NS)
            if hp is None:
                hp = etree.SubElement(pm, f"{{{NS['wpml']}}}waypointHeadingParam")
            for child in list(hp):
                hp.remove(child)
            etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingMode").text = "fixed"
            etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingAngle").text = str(int(yaw_angle))
            etree.SubElement(hp, f"{{{NS['wpml']}}}waypointPoiPoint").text = "0.000000,0.000000,0.000000"
            etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingPathMode").text = "followBadArc"
            etree.SubElement(hp, f"{{{NS['wpml']}}}waypointHeadingPoiIndex").text = "0"

    # --- 写真モード ----------------------------------------------------------
    if do_photo and not do_video:
        placemarks = [
            p for p in tree.findall(".//kml:Placemark", NS)
            if p.find("wpml:index", NS) is not None
        ]
        for pm in placemarks:
            idx = int(pm.find("wpml:index", NS).text)
            ag = etree.SubElement(pm, f"{{{NS['wpml']}}}actionGroup")
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupId").text = str(idx)
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupStartIndex").text = str(idx)
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupEndIndex").text = str(idx)
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupMode").text = "sequence"
            trg = etree.SubElement(ag, f"{{{NS['wpml']}}}actionTrigger")
            etree.SubElement(trg, f"{{{NS['wpml']}}}actionTriggerType").text = "reachPoint"

            # 1) rotateYaw アクション
            yaw_target = None
            if yaw_fix:
                if yaw_mode == "original" and original_angles and idx in original_angles:
                    yaw_target = original_angles[idx].get("heading")
                elif yaw_angle is not None:
                    yaw_target = yaw_angle
            if yaw_target is not None:
                yaw_action = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                etree.SubElement(yaw_action, f"{{{NS['wpml']}}}actionId").text = "0"
                etree.SubElement(yaw_action, f"{{{NS['wpml']}}}actionActuatorFunc").text = "rotateYaw"
                yaw_param = etree.SubElement(yaw_action, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                etree.SubElement(yaw_param, f"{{{NS['wpml']}}}aircraftHeading").text = str(int(yaw_target))
                etree.SubElement(yaw_param, f"{{{NS['wpml']}}}aircraftPathMode").text = "counterClockwise"

            # 2) ジンバルピッチ アクション
            pitch_target = None
            if do_gimbal:
                if gimbal_pitch_mode == "original" and original_angles and idx in original_angles:
                    pitch_target = original_angles[idx].get("pitch")
                elif gimbal_pitch_angle is not None:
                    pitch_target = gimbal_pitch_angle
            if pitch_target is not None:
                gp_action = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                etree.SubElement(gp_action, f"{{{NS['wpml']}}}actionId").text = "0"
                etree.SubElement(gp_action, f"{{{NS['wpml']}}}actionActuatorFunc").text = "gimbalRotate"
                param = etree.SubElement(gp_action, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalRotateMode").text = "absoluteAngle"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalPitchRotateEnable").text = "1"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalPitchRotateAngle").text = str(int(pitch_target))
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalRollRotateEnable").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalRollRotateAngle").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalYawRotateEnable").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalYawRotateAngle").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalRotateTimeEnable").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalRotateTime").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"

            # 3) takePhoto
            act = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
            etree.SubElement(act, f"{{{NS['wpml']}}}actionId").text = "0"
            etree.SubElement(act, f"{{{NS['wpml']}}}actionActuatorFunc").text = "takePhoto"
            param = etree.SubElement(act, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(param, f"{{{NS['wpml']}}}fileSuffix").text = f"ウェイポイント{idx}"
            etree.SubElement(param, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"
            etree.SubElement(param, f"{{{NS['wpml']}}}useGlobalPayloadLensIndex").text = "1"

            # 4) ホバリング
            if hover_time > 0:
                hover_action = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                etree.SubElement(hover_action, f"{{{NS['wpml']}}}actionId").text = "0"
                etree.SubElement(hover_action, f"{{{NS['wpml']}}}actionActuatorFunc").text = "hover"
                hover_param = etree.SubElement(hover_action, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                etree.SubElement(hover_param, f"{{{NS['wpml']}}}hoverTime").text = str(int(hover_time))

    # --- 動画モード ----------------------------------------------------------
    if do_video:
        pms = sorted(
            [p for p in tree.findall(".//kml:Placemark", NS)
             if p.find("wpml:index", NS) is not None],
            key=lambda p: int(p.find("wpml:index", NS).text)
        )
        if pms:
            first, last = pms[0], pms[-1]
            first_idx = int(first.find("wpml:index", NS).text)
            # 録画開始アクション
            ag = etree.SubElement(first, f"{{{NS['wpml']}}}actionGroup")
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupId").text = "0"
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupStartIndex").text = "0"
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupEndIndex").text = "0"
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupMode").text = "sequence"
            trg = etree.SubElement(ag, f"{{{NS['wpml']}}}actionTrigger")
            etree.SubElement(trg, f"{{{NS['wpml']}}}actionTriggerType").text = "reachPoint"

            # rotateYaw & gimbalRotate アクション
            yaw_target = None
            if yaw_fix:
                if yaw_mode == "original" and original_angles and first_idx in original_angles:
                    yaw_target = original_angles[first_idx].get("heading")
                elif yaw_angle is not None:
                    yaw_target = yaw_angle
            if yaw_target is not None:
                yaw_action = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                etree.SubElement(yaw_action, f"{{{NS['wpml']}}}actionId").text = "0"
                etree.SubElement(yaw_action, f"{{{NS['wpml']}}}actionActuatorFunc").text = "rotateYaw"
                yaw_param = etree.SubElement(yaw_action, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                etree.SubElement(yaw_param, f"{{{NS['wpml']}}}aircraftHeading").text = str(int(yaw_target))
                etree.SubElement(yaw_param, f"{{{NS['wpml']}}}aircraftPathMode").text = "counterClockwise"

            pitch_target = None
            if do_gimbal:
                if gimbal_pitch_mode == "original" and original_angles and first_idx in original_angles:
                    pitch_target = original_angles[first_idx].get("pitch")
                elif gimbal_pitch_angle is not None:
                    pitch_target = gimbal_pitch_angle
            if pitch_target is not None:
                gp_action = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                etree.SubElement(gp_action, f"{{{NS['wpml']}}}actionId").text = "0"
                etree.SubElement(gp_action, f"{{{NS['wpml']}}}actionActuatorFunc").text = "gimbalRotate"
                param = etree.SubElement(gp_action, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalRotateMode").text = "absoluteAngle"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalPitchRotateEnable").text = "1"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalPitchRotateAngle").text = str(int(pitch_target))
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalRollRotateEnable").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalRollRotateAngle").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalYawRotateEnable").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalYawRotateAngle").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalRotateTimeEnable").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}gimbalRotateTime").text = "0"
                etree.SubElement(param, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"

            # startRecord
            acs = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
            etree.SubElement(acs, f"{{{NS['wpml']}}}actionId").text = "0"
            etree.SubElement(acs, f"{{{NS['wpml']}}}actionActuatorFunc").text = "startRecord"
            p = etree.SubElement(acs, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(p, f"{{{NS['wpml']}}}fileSuffix").text = video_suffix
            etree.SubElement(p, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"

            # stopRecord（最終ポイント）
            ag2 = etree.SubElement(last, f"{{{NS['wpml']}}}actionGroup")
            idx_end = len(pms) - 1
            etree.SubElement(ag2, f"{{{NS['wpml']}}}actionGroupId").text = str(idx_end)
            etree.SubElement(ag2, f"{{{NS['wpml']}}}actionGroupStartIndex").text = str(idx_end)
            etree.SubElement(ag2, f"{{{NS['wpml']}}}actionGroupEndIndex").text = str(idx_end)
            etree.SubElement(ag2, f"{{{NS['wpml']}}}actionGroupMode").text = "sequence"
            trg2 = etree.SubElement(ag2, f"{{{NS['wpml']}}}actionTrigger")
            etree.SubElement(trg2, f"{{{NS['wpml']}}}actionTriggerType").text = "reachPoint"
            st = etree.SubElement(ag2, f"{{{NS['wpml']}}}action")
            etree.SubElement(st, f"{{{NS['wpml']}}}actionId").text = "0"
            etree.SubElement(st, f"{{{NS['wpml']}}}actionActuatorFunc").text = "stopRecord"
            sparam = etree.SubElement(st, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(sparam, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"

# --- KMZ 一括処理 -----------------------------------------------------------
def process_kmz(path, offset, do_photo, do_video, video_suffix,
                do_gimbal, gimbal_pitch_angle, gimbal_pitch_mode,
                yaw_fix, yaw_angle, yaw_mode, speed, sensor_modes, hover_time,
                coordinate_deviation, log):
    try:
        log.insert(tk.END, f"Extracting {os.path.basename(path)}...\n")
        wd = extract_kmz(path)
        kmls = glob.glob(os.path.join(wd, "**", "template.kml"), recursive=True)
        if not kmls:
            raise FileNotFoundError("template.kml が見つかりませんでした。")
        out_root, outdir = prepare_output_dirs(path, offset)
        original_angles = None
        for kml in kmls:
            log.insert(tk.END, f"Converting {os.path.basename(kml)}...\n")
            parser = etree.XMLParser(remove_blank_text=True)
            tree = etree.parse(kml, parser)
            if yaw_mode == "original" or gimbal_pitch_mode == "original":
                original_angles = extract_original_gimbal_angles(tree)
                if original_angles:
                    log.insert(tk.END, f"Found original angles for {len(original_angles)} waypoints\n")
                else:
                    log.insert(tk.END, "No original gimbal data found\n")
            convert_kml(tree, offset, do_photo, do_video, video_suffix,
                        do_gimbal, gimbal_pitch_angle, gimbal_pitch_mode,
                        yaw_fix, yaw_angle, yaw_mode, speed, sensor_modes, hover_time,
                        coordinate_deviation, original_angles)
            out_path = os.path.join(outdir, os.path.basename(kml))
            tree.write(out_path, encoding="utf-8",
                       pretty_print=True, xml_declaration=True)
        for name in ["res"]:
            srcs = glob.glob(os.path.join(wd, "**", name), recursive=True)
            if srcs:
                src = srcs[0]
                dst = os.path.join(outdir, os.path.basename(src))
                if os.path.isdir(src):
                    if os.path.exists(dst):
                        shutil.rmtree(dst)
                    shutil.copytree(src, dst)
                else:
                    shutil.copy2(src, dst)
        out_kmz = repackage_to_kmz(out_root, path)
        log.insert(tk.END, f"Saved: {out_kmz}\nFinished\n\n")
        messagebox.showinfo("完了", f"変換完了:\n{out_kmz}")
    except Exception as e:
        messagebox.showerror("エラー", str(e))
        log.insert(tk.END, f"Error: {e}\n\n")
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
        self.hc.grid(row=0, column=1, padx=5, columnspan=2, sticky="w")
        self.hc.bind("<<ComboboxSelected>>", self.on_height_change)
        self.he = ttk.Entry(self, width=10, state="disabled")
        self.he.grid(row=0, column=3, padx=5)

        # 速度
        ttk.Label(self, text="速度 (1–15 m/s):").grid(row=1, column=0, sticky="w", pady=5)
        self.sp = tk.IntVar(value=15)
        ttk.Spinbox(self, from_=1, to=15, textvariable=self.sp, width=5).grid(row=1, column=1, columnspan=2, sticky="w")

        # 撮影設定（排他制御付き）
        self.ph = tk.BooleanVar(value=False)
        self.vd = tk.BooleanVar(value=False)
        self.cb_photo = ttk.Checkbutton(self, text="写真撮影", variable=self.ph, command=self.on_photo_toggle)
        self.cb_photo.grid(row=2, column=0, sticky="w")
        self.cb_video = ttk.Checkbutton(self, text="動画撮影", variable=self.vd, command=self.on_video_toggle)
        self.cb_video.grid(row=2, column=1, sticky="w")
        self.vd_suffix_label = ttk.Label(self, text="動画ファイル名:")
        self.vd_suffix_var = tk.StringVar(value="video_01")
        self.vd_suffix_entry = ttk.Entry(self, textvariable=self.vd_suffix_var, width=20)

        # センサーモード
        ttk.Label(self, text="センサー選択:").grid(row=3, column=0, sticky="w")
        self.sm_vars = {m: tk.BooleanVar(value=False) for m in SENSOR_MODES}
        for i, m in enumerate(SENSOR_MODES):
            ttk.Checkbutton(self, text=m, variable=self.sm_vars[m]).grid(row=3, column=1 + i, sticky="w")

        # ジンバルピッチ
        self.gp_var = tk.BooleanVar(value=False)
        self.gp_cb = ttk.Checkbutton(self, text="ジンバルピッチ", variable=self.gp_var, command=self.update_gimbal_pitch)
        self.gp_cb.grid(row=4, column=0, sticky="w", pady=5)
        self.gp_mode = ttk.Combobox(self, values=list(GIMBAL_PITCH_OPTIONS), state="readonly", width=15)
        self.gp_mode.bind("<<ComboboxSelected>>", self.update_gimbal_pitch)
        self.gp_entry = ttk.Entry(self, width=8, state="disabled")

        # ヨー固定
        self.yf = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="ヨー固定", variable=self.yf, command=self.update_yaw).grid(row=5, column=0, sticky="w")
        self.yc = ttk.Combobox(self, values=list(YAW_OPTIONS), state="readonly", width=15)
        self.yc.bind("<<ComboboxSelected>>", self.update_yaw)
        self.ye = ttk.Entry(self, width=8, state="disabled")

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
        self.ref_point_combo = ttk.Combobox(self, textvariable=self.ref_point_var,
                                            values=list(REFERENCE_POINTS.keys()),
                                            state="readonly", width=10)
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
        self.update_gimbal_pitch()
        self.update_yaw()
        self.update_hover()
        self.update_deviation()

    # 高度選択時
    def on_height_change(self, event=None):
        if HEIGHT_OPTIONS.get(self.hc.get()) == "custom":
            self.he.config(state="normal")
            self.he.delete(0, tk.END)
            self.he.focus()
        else:
            self.he.config(state="disabled")
            self.he.delete(0, tk.END)

    # 写真/動画 排他制御
    def on_photo_toggle(self):
        if self.ph.get():
            # 写真ON → 動画OFF
            self.vd.set(False)
        self.update_ctrl()

    def on_video_toggle(self):
        if self.vd.get():
            # 動画ON → 写真OFF
            self.ph.set(False)
        self.update_ctrl()

    # 撮影設定表示更新
    def update_ctrl(self):
        if self.vd.get():
            self.vd_suffix_label.grid(row=2, column=2, sticky="e", padx=(10, 2))
            self.vd_suffix_entry.grid(row=2, column=3, sticky="w")
        else:
            self.vd_suffix_label.grid_forget()
            self.vd_suffix_entry.grid_forget()

    # 以下、元機能のUI更新ハンドラを維持
    def update_gimbal_pitch(self, event=None):
        if self.gp_var.get():
            self.gp_mode.grid(row=4, column=1, padx=5, columnspan=2, sticky="w")
            if not self.gp_mode.get():
                self.gp_mode.set(next(iter(GIMBAL_PITCH_OPTIONS)))
            if self.gp_mode.get() == "手動入力":
                self.gp_entry.config(state="normal")
                self.gp_entry.grid(row=4, column=3)
            else:
                self.gp_entry.config(state="disabled")
                self.gp_entry.grid_forget()
        else:
            self.gp_mode.grid_forget()
            self.gp_entry.grid_forget()

    def update_yaw(self, event=None):
        if self.yf.get():
            self.yc.grid(row=5, column=1, padx=5, columnspan=2, sticky="w")
            if not self.yc.get():
                self.yc.set(next(iter(YAW_OPTIONS)))
            if self.yc.get() == "手動入力":
                self.ye.config(state="normal")
                self.ye.grid(row=5, column=3)
            else:
                self.ye.config(state="disabled")
                self.ye.grid_forget()
        else:
            self.yc.grid_forget()
            self.ye.grid_forget()

    def update_hover(self):
        if self.hv.get():
            self.hover_time_label.grid(row=6, column=1, sticky="e", padx=(10, 2))
            self.hover_time_entry.grid(row=6, column=2, sticky="w")
        else:
            self.hover_time_label.grid_forget()
            self.hover_time_entry.grid_forget()

    def update_deviation(self):
        if self.dc.get():
            self.ref_point_label.grid(row=7, column=1, sticky="e", padx=(10, 2))
            self.ref_point_combo.grid(row=7, column=2, sticky="w")
            self.today_coords_label.grid(row=8, column=0, sticky="w", pady=5)
            self.today_lng_entry.grid(row=8, column=1, padx=2, sticky="w")
            self.today_lat_entry.grid(row=8, column=2, padx=2, sticky="w")
            self.today_alt_entry.grid(row=8, column=3, padx=2, sticky="w")
            self.copy_button.grid(row=8, column=4, padx=5, sticky="w")
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
            ref_point = self.ref_point_var.get()
            ref_coords = REFERENCE_POINTS[ref_point]
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            copy_text = f"{current_time},{ref_coords[0]},{ref_coords[1]},{ref_coords[2]}"
            pyperclip.copy(copy_text)
            messagebox.showinfo("コピー完了", f"クリップボードにコピーしました:\n{copy_text}")
        except Exception as e:
            messagebox.showerror("エラー", f"コピーに失敗しました: {e}")

    # パラメータ収集（元コード）
    def get_params(self):
        offset = 0.0
        v = HEIGHT_OPTIONS.get(self.hc.get())
        if v == "custom":
            try:
                offset = float(self.he.get())
            except:
                pass
        else:
            offset = float(v)

        yaw_angle = None
        yaw_mode = "none"
        if self.yf.get():
            yval = YAW_OPTIONS.get(self.yc.get())
            if yval == "original":
                yaw_mode = "original"
            elif yval == "custom":
                yaw_mode = "fixed"
                try:
                    yaw_angle = float(self.ye.get())
                except:
                    pass
            else:
                yaw_mode = "fixed"
                yaw_angle = float(yval)

        gimbal_pitch_angle = None
        gimbal_pitch_mode = "none"
        do_gimbal = self.gp_var.get()
        if do_gimbal:
            gp_val = GIMBAL_PITCH_OPTIONS.get(self.gp_mode.get())
            if gp_val == "original":
                gimbal_pitch_mode = "original"
            elif gp_val == "custom":
                gimbal_pitch_mode = "fixed"
                try:
                    gimbal_pitch_angle = float(self.gp_entry.get())
                except:
                    pass
            else:
                gimbal_pitch_mode = "fixed"
                gimbal_pitch_angle = float(gp_val)

        hover_time = 0
        if self.hv.get():
            try:
                hover_time = max(0, float(self.hover_time_var.get()))
            except:
                hover_time = 2.0

        coordinate_deviation = None
        if self.dc.get():
            try:
                ref_coords = REFERENCE_POINTS[self.ref_point_var.get()]
                current_coords = (float(self.today_lng_var.get()),
                                  float(self.today_lat_var.get()),
                                  float(self.today_alt_var.get()))
                deviation = calculate_deviation(ref_coords, current_coords)
                is_safe, error_msg = check_deviation_safety(deviation)
                if not is_safe:
                    raise ValueError(error_msg)
                coordinate_deviation = deviation
            except Exception as e:
                messagebox.showerror("偏差補正エラー", str(e))
                return None

        return {
            "offset": offset,
            "do_photo": self.ph.get(),
            "do_video": self.vd.get(),
            "video_suffix": self.vd_suffix_var.get(),
            "do_gimbal": do_gimbal,
            "gimbal_pitch_angle": gimbal_pitch_angle,
            "gimbal_pitch_mode": gimbal_pitch_mode,
            "yaw_fix": self.yf.get(),
            "yaw_angle": yaw_angle,
            "yaw_mode": yaw_mode,
            "speed": max(1, min(15, self.sp.get())),
            "sensor_modes": [m for m, var in self.sm_vars.items() if var.get()],
            "hover_time": hover_time,
            "coordinate_deviation": coordinate_deviation
        }

# --- エントリポイント -------------------------------------------------------
def main():
    root = TkinterDnD.Tk()
    root.title("ATL→ASL 変換＋撮影制御ツール (ver. GUI53 – 排他制御版)")
    root.geometry("800x780")
    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)
    app = AppGUI(frm)
    app.pack(fill="x", pady=(0, 10))

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
            messagebox.showwarning("警告", ".kmz ファイルのみ対応しています。")
            return
        params = app.get_params()
        if params is None:
            return
        params.update({"path": path, "log": log})
        threading.Thread(target=process_kmz, kwargs=params, daemon=True).start()

    drop.dnd_bind("<<Drop>>", on_drop)
    root.mainloop()

if __name__ == "__main__":
    main()
