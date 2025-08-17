#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
convert_height_gui_asl.py (ver. GUI32 - ヨー角制御最適化版)

変更点
・写真撮影+ヨー角固定時の効率的なアクション配置を実装
・不要なアクションの削除とより適切なアクション順序
・orientedShootの適切な使用
・rotateYawパラメータの修正（aircraftHeading使用）
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
    "1Q: 81.00°": 81.00,
    "2Q: 96.92°": 96.92,
    "4Q: 87.31°": 87.31,
    "手動入力": "custom"
}

SENSOR_MODES = ["Wide", "Zoom", "IR"]

# 基準ポイント座標
REFERENCE_POINTS = {
    "本部": (136.5559522506280, 36.0729517605894, 612.2),
    "烏帽子": (136.560000000000, 36.075000000000, 962.02)  # 仮の座標
}

# 20m偏差の閾値（緯度経度）
DEVIATION_THRESHOLD = {
    "lat": 0.00018,  # 約20m in latitude
    "lng": 0.00022,  # 約20m in longitude at 36°N
    "alt": 20.0      # 20m in altitude
}

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
    """基準座標と現在座標から偏差を計算"""
    ref_lng, ref_lat, ref_alt = reference_coords
    cur_lng, cur_lat, cur_alt = current_coords
    dev_lng = cur_lng - ref_lng
    dev_lat = cur_lat - ref_lat
    dev_alt = cur_alt - ref_alt
    return dev_lng, dev_lat, dev_alt

def check_deviation_safety(deviation):
    """偏差が安全範囲内かチェック"""
    dev_lng, dev_lat, dev_alt = deviation
    if (abs(dev_lng) > DEVIATION_THRESHOLD["lng"] or
        abs(dev_lat) > DEVIATION_THRESHOLD["lat"] or
        abs(dev_alt) > DEVIATION_THRESHOLD["alt"]):
        return False, f"偏差が20mを超えています:\n経度偏差: {dev_lng:.8f}°\n緯度偏差: {dev_lat:.8f}°\n標高偏差: {dev_alt:.2f}m"
    return True, None

# --- KML 変換 ---------------------------------------------------------------
def convert_kml(tree, offset, do_photo, do_video, video_suffix,
                do_gimbal, yaw_fix, yaw_angle, speed, sensor_modes, hover_time,
                coordinate_deviation=None):
    
    def _max_id(root, xp):
        ids = [int(e.text) for e in root.findall(xp, NS)
               if e.text and e.text.isdigit()]
        return max(ids) if ids else -1

    # 1) 座標偏差補正（新機能）
    if coordinate_deviation:
        dev_lng, dev_lat, dev_alt = coordinate_deviation
        # 各ウェイポイントの座標を補正
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

    # 2) 高度補正＋EGM96
    for pm in tree.findall(".//kml:Placemark", NS):
        for tag in ("height", "ellipsoidHeight"):
            el = pm.find(f"wpml:{tag}", NS)
            if el is not None and el.text:
                try:
                    current_height = float(el.text) + offset
                    # 座標偏差補正がある場合は標高偏差も適用
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

    # 3) 速度設定
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

    # 4) 既存アクション削除
    for ag in tree.findall(".//wpml:actionGroup", NS):
        for act in list(ag.findall("wpml:action", NS)):
            f = act.find("wpml:actionActuatorFunc", NS)
            if f is not None and f.text in (
                "orientedShoot", "startRecord", "stopRecord", "takePhoto",
                "gimbalRotate", "selectWide", "selectZoom", "selectIR", "hover", "rotateYaw"
            ):
                ag.remove(act)

    pp = tree.find(".//wpml:payloadParam", NS)
    if pp is not None:
        img = pp.find("wpml:imageFormat", NS)
        if img is not None:
            pp.remove(img)

    # 5) センサー選択
    if sensor_modes:
        if pp is None:
            fld = tree.find(".//kml:Folder", NS)
            pp = etree.SubElement(fld, f"{{{NS['wpml']}}}payloadParam")
            etree.SubElement(pp, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"
        fmt = ",".join(m.lower() for m in sensor_modes)
        etree.SubElement(pp, f"{{{NS['wpml']}}}imageFormat").text = fmt

    # 6) ヨー角グローバル設定
    if yaw_fix and yaw_angle is not None:
        for hp in tree.findall(".//wpml:globalWaypointHeadingParam", NS):
            m = hp.find("wpml:waypointHeadingMode", NS)
            a = hp.find("wpml:waypointHeadingAngle", NS)
            if m is not None: 
                m.text = "fixed"
            if a is not None: 
                a.text = str(yaw_angle)

    # 7) 写真撮影（最適化版）
    if do_photo and not do_video:
        pms = [
            p for p in tree.findall(".//kml:Placemark", NS)
            if p.find("wpml:index", NS) is not None
        ]
        pms.sort(key=lambda p: int(p.find("wpml:index", NS).text))

        base_action_group_id = _max_id(tree.getroot(), ".//wpml:actionGroupId")

        for i, pm in enumerate(pms):
            idx = int(pm.find("wpml:index", NS).text)
            
            # actionGroupを作成
            ag = etree.SubElement(pm, f"{{{NS['wpml']}}}actionGroup")
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupId").text = str(base_action_group_id + i + 1)
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupStartIndex").text = str(idx)
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupEndIndex").text = str(idx)
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupMode").text = "sequence"
            trg = etree.SubElement(ag, f"{{{NS['wpml']}}}actionTrigger")
            etree.SubElement(trg, f"{{{NS['wpml']}}}actionTriggerType").text = "reachPoint"

            action_id = 0
            
            # ヨー角制御の最適化実装
            if yaw_fix and yaw_angle is not None:
                # 現在のウェイポイントの設定を確認
                wp_heading_param = pm.find("wpml:waypointHeadingParam", NS)
                current_angle = None
                if wp_heading_param is not None:
                    angle_elem = wp_heading_param.find("wpml:waypointHeadingAngle", NS)
                    if angle_elem is not None and angle_elem.text:
                        try:
                            current_angle = float(angle_elem.text)
                        except:
                            pass
                
                # 角度が異なる場合、またはウェイポイント設定が存在しない場合のみrotateYawを追加
                need_rotate_yaw = (current_angle is None or abs(current_angle - yaw_angle) > 0.5)
                
                if need_rotate_yaw:
                    # rotateYawアクション
                    yaw_action = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                    etree.SubElement(yaw_action, f"{{{NS['wpml']}}}actionId").text = str(action_id)
                    etree.SubElement(yaw_action, f"{{{NS['wpml']}}}actionActuatorFunc").text = "rotateYaw"
                    yaw_param = etree.SubElement(yaw_action, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                    etree.SubElement(yaw_param, f"{{{NS['wpml']}}}aircraftHeading").text = str(yaw_angle)
                    etree.SubElement(yaw_param, f"{{{NS['wpml']}}}aircraftPathMode").text = "counterClockwise"
                    action_id += 1
                    
                    # 個別のウェイポイント設定を更新
                    etree.SubElement(pm, f"{{{NS['wpml']}}}useGlobalSpeed").text = "0"
                    
                    # waypointHeadingParamを更新または作成
                    if wp_heading_param is None:
                        wp_heading_param = etree.SubElement(pm, f"{{{NS['wpml']}}}waypointHeadingParam")
                    
                    # 設定を更新
                    mode_elem = wp_heading_param.find("wpml:waypointHeadingMode", NS)
                    if mode_elem is None:
                        etree.SubElement(wp_heading_param, f"{{{NS['wpml']}}}waypointHeadingMode").text = "fixed"
                    else:
                        mode_elem.text = "fixed"
                    
                    angle_elem = wp_heading_param.find("wpml:waypointHeadingAngle", NS)
                    if angle_elem is None:
                        etree.SubElement(wp_heading_param, f"{{{NS['wpml']}}}waypointHeadingAngle").text = str(yaw_angle)
                    else:
                        angle_elem.text = str(yaw_angle)
                        
                    # その他の必要な設定
                    poi_elem = wp_heading_param.find("wpml:waypointPoiPoint", NS)
                    if poi_elem is None:
                        etree.SubElement(wp_heading_param, f"{{{NS['wpml']}}}waypointPoiPoint").text = "0.000000,0.000000,0.000000"
                    
                    path_mode_elem = wp_heading_param.find("wpml:waypointHeadingPathMode", NS)
                    if path_mode_elem is None:
                        etree.SubElement(wp_heading_param, f"{{{NS['wpml']}}}waypointHeadingPathMode").text = "followBadArc"
                    
                    poi_index_elem = wp_heading_param.find("wpml:waypointHeadingPoiIndex", NS)
                    if poi_index_elem is None:
                        etree.SubElement(wp_heading_param, f"{{{NS['wpml']}}}waypointHeadingPoiIndex").text = "0"
                
                else:
                    # rotateYawが不要な場合はグローバル設定を使用
                    etree.SubElement(pm, f"{{{NS['wpml']}}}useGlobalHeadingParam").text = "1"
                    etree.SubElement(pm, f"{{{NS['wpml']}}}useGlobalSpeed").text = "0"

            # ジンバル制御（写真撮影時のみ適用）
            if do_gimbal:
                gimbal_action = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                etree.SubElement(gimbal_action, f"{{{NS['wpml']}}}actionId").text = str(action_id)
                etree.SubElement(gimbal_action, f"{{{NS['wpml']}}}actionActuatorFunc").text = "gimbalRotate"
                param = etree.SubElement(gimbal_action, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                param_map = {
                    "gimbalRotateMode": "absoluteAngle",
                    "gimbalPitchRotateEnable": 1,
                    "gimbalPitchRotateAngle": -90,
                    "gimbalRollRotateEnable": 0,
                    "gimbalRollRotateAngle": 0,
                    "gimbalYawRotateEnable": 0,
                    "gimbalYawRotateAngle": 0,
                    "gimbalRotateTimeEnable": 0,
                    "gimbalRotateTime": 0,
                    "payloadPositionIndex": 0
                }
                for k, v in param_map.items():
                    etree.SubElement(param, f"{{{NS['wpml']}}}{k}").text = str(v)
                action_id += 1

            # ホバリング（ユーザーが選択した場合のみ）
            if hover_time > 0:
                hover_action = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                etree.SubElement(hover_action, f"{{{NS['wpml']}}}actionId").text = str(action_id)
                etree.SubElement(hover_action, f"{{{NS['wpml']}}}actionActuatorFunc").text = "hover"
                hover_param = etree.SubElement(hover_action, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                etree.SubElement(hover_param, f"{{{NS['wpml']}}}hoverTime").text = str(int(hover_time))
                action_id += 1

            # 写真撮影アクション（orientedShootを使用）
            shoot_action = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
            etree.SubElement(shoot_action, f"{{{NS['wpml']}}}actionId").text = str(action_id)
            etree.SubElement(shoot_action, f"{{{NS['wpml']}}}actionActuatorFunc").text = "orientedShoot"
            etree.SubElement(shoot_action, f"{{{NS['wpml']}}}actionActuatorFuncParam")

            # 共通設定
            etree.SubElement(pm, f"{{{NS['wpml']}}}useGlobalTurnParam").text = "1"
            etree.SubElement(pm, f"{{{NS['wpml']}}}useStraightLine").text = "0"
            etree.SubElement(pm, f"{{{NS['wpml']}}}isRisky").text = "0"

    # 8) 動画撮影（最適化）
    if do_video:
        pms = sorted([
            p for p in tree.findall(".//kml:Placemark", NS)
            if p.find("wpml:index", NS) is not None
        ], key=lambda p: int(p.find("wpml:index", NS).text))
        
        if pms:
            first, last = pms[0], pms[-1]
            base = _max_id(tree.getroot(), ".//wpml:actionGroupId")
            
            # 開始ポイントの処理
            ag = etree.SubElement(first, f"{{{NS['wpml']}}}actionGroup")
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupId").text = str(base + 1)
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupStartIndex").text = first.find("wpml:index", NS).text
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupEndIndex").text = first.find("wpml:index", NS).text
            etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupMode").text = "sequence"
            trg = etree.SubElement(ag, f"{{{NS['wpml']}}}actionTrigger")
            etree.SubElement(trg, f"{{{NS['wpml']}}}actionTriggerType").text = "reachPoint"
            
            nid = 0
            
            # ジンバル回転（オプション化）
            if do_gimbal:
                gimbal = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                etree.SubElement(gimbal, f"{{{NS['wpml']}}}actionId").text = str(nid)
                etree.SubElement(gimbal, f"{{{NS['wpml']}}}actionActuatorFunc").text = "gimbalRotate"
                param = etree.SubElement(gimbal, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                param_map = {
                    "gimbalRotateMode": "absoluteAngle",
                    "gimbalPitchRotateEnable": 1,
                    "gimbalPitchRotateAngle": -90,
                    "gimbalRollRotateEnable": 0,
                    "gimbalRollRotateAngle": 0,
                    "gimbalYawRotateEnable": 0,
                    "gimbalYawRotateAngle": 0,
                    "gimbalRotateTimeEnable": 0,
                    "gimbalRotateTime": 0,
                    "payloadPositionIndex": 0
                }
                for k, v in param_map.items():
                    etree.SubElement(param, f"{{{NS['wpml']}}}{k}").text = str(v)
                nid += 1
            
            # 録画開始
            acs = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
            etree.SubElement(acs, f"{{{NS['wpml']}}}actionId").text = str(nid)
            etree.SubElement(acs, f"{{{NS['wpml']}}}actionActuatorFunc").text = "startRecord"
            p = etree.SubElement(acs, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(p, f"{{{NS['wpml']}}}fileSuffix").text = video_suffix
            etree.SubElement(p, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"
            
            # 録画停止（最終ポイント）
            ag2 = etree.SubElement(last, f"{{{NS['wpml']}}}actionGroup")
            etree.SubElement(ag2, f"{{{NS['wpml']}}}actionGroupId").text = str(base + 2)
            etree.SubElement(ag2, f"{{{NS['wpml']}}}actionGroupStartIndex").text = last.find("wpml:index", NS).text
            etree.SubElement(ag2, f"{{{NS['wpml']}}}actionGroupEndIndex").text = last.find("wpml:index", NS).text
            etree.SubElement(ag2, f"{{{NS['wpml']}}}actionGroupMode").text = "sequence"
            trg2 = etree.SubElement(ag2, f"{{{NS['wpml']}}}actionTrigger")
            etree.SubElement(trg2, f"{{{NS['wpml']}}}actionTriggerType").text = "reachPoint"
            
            st = etree.SubElement(ag2, f"{{{NS['wpml']}}}action")
            etree.SubElement(st, f"{{{NS['wpml']}}}actionId").text = "0"
            etree.SubElement(st, f"{{{NS['wpml']}}}actionActuatorFunc").text = "stopRecord"
            sparam = etree.SubElement(st, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(sparam, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"

    # 9) 空の actionGroup 削除
    for ag in tree.findall(".//wpml:actionGroup", NS):
        if not ag.findall("wpml:action", NS):
            ag.getparent().remove(ag)

def process_kmz(path, offset, do_photo, do_video, video_suffix,
                do_gimbal, yaw_fix, yaw_angle, speed, sensor_modes, hover_time,
                coordinate_deviation, log):
    try:
        log.insert(tk.END, f"Extracting {os.path.basename(path)}...\n")
        wd = extract_kmz(path)
        kmls = glob.glob(os.path.join(wd, "**", "template.kml"), recursive=True)
        if not kmls:
            raise FileNotFoundError("template.kml が見つかりませんでした。")
        
        out_root, outdir = prepare_output_dirs(path, offset)
        
        for kml in kmls:
            log.insert(tk.END, f"Converting {os.path.basename(kml)}...\n")
            parser = etree.XMLParser(remove_blank_text=True)
            tree = etree.parse(kml, parser)
            convert_kml(tree, offset, do_photo, do_video, video_suffix,
                       do_gimbal, yaw_fix, yaw_angle, speed, sensor_modes, hover_time,
                       coordinate_deviation)
            out_path = os.path.join(outdir, os.path.basename(kml))
            tree.write(out_path, encoding="utf-8",
                      pretty_print=True, xml_declaration=True)
        
        # "waylines.wpml" はコピーせず、"res" のみ
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

class AppGUI(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        
        # 基準高度設定
        ttk.Label(self, text="基準高度:").grid(row=0, column=0, sticky="w")
        self.hc = ttk.Combobox(self, values=list(HEIGHT_OPTIONS), state="readonly", width=20)
        self.hc.set(next(iter(HEIGHT_OPTIONS)))
        self.hc.grid(row=0, column=1, padx=5, columnspan=2, sticky="w")
        self.hc.bind("<<ComboboxSelected>>", self.on_height_change)
        self.he = ttk.Entry(self, width=10, state="disabled")
        self.he.grid(row=0, column=3, padx=5)
        
        # 速度設定
        ttk.Label(self, text="速度 (1–15 m/s):").grid(row=1, column=0, sticky="w", pady=5)
        self.sp = tk.IntVar(value=15)
        ttk.Spinbox(self, from_=1, to=15, textvariable=self.sp, width=5).grid(row=1, column=1, columnspan=2, sticky="w")
        
        # 撮影設定
        self.ph = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="写真撮影", variable=self.ph, command=self.update_ctrl).grid(row=2, column=0, sticky="w")
        self.vd = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="動画撮影", variable=self.vd, command=self.update_ctrl).grid(row=2, column=1, sticky="w")
        
        self.vd_suffix_label = ttk.Label(self, text="動画ファイル名:")
        self.vd_suffix_var = tk.StringVar(value="video_01")
        self.vd_suffix_entry = ttk.Entry(self, textvariable=self.vd_suffix_var, width=20)
        
        # センサー選択
        ttk.Label(self, text="センサー選択:").grid(row=3, column=0, sticky="w")
        self.sm_vars = {m: tk.BooleanVar(value=False) for m in SENSOR_MODES}
        for i, m in enumerate(SENSOR_MODES):
            ttk.Checkbutton(self, text=m, variable=self.sm_vars[m]).grid(row=3, column=1+i, sticky="w")
        
        # ジンバル制御
        self.gm = tk.BooleanVar(value=True)
        self.gc = ttk.Checkbutton(self, text="ジンバル制御", variable=self.gm)
        self.gc.grid(row=4, column=0, sticky="w", pady=5)
        
        # ヨー固定設定
        self.yf = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="ヨー固定", variable=self.yf, command=self.update_yaw).grid(row=5, column=0, sticky="w")
        self.yc = ttk.Combobox(self, values=list(YAW_OPTIONS), state="readonly", width=15)
        self.yc.bind("<<ComboboxSelected>>", self.update_yaw)
        self.ye = ttk.Entry(self, width=8, state="disabled")
        
        # ホバリング設定
        self.hv = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="ホバリング", variable=self.hv, command=self.update_hover).grid(row=6, column=0, sticky="w", pady=5)
        self.hover_time_label = ttk.Label(self, text="ホバリング時間 (秒):")
        self.hover_time_var = tk.StringVar(value="2")
        self.hover_time_entry = ttk.Entry(self, textvariable=self.hover_time_var, width=8)
        
        # 偏差補正設定（新機能）
        self.dc = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="偏差補正", variable=self.dc, command=self.update_deviation).grid(row=7, column=0, sticky="w", pady=5)
        
        # 基準位置選択
        self.ref_point_label = ttk.Label(self, text="基準位置:")
        self.ref_point_var = tk.StringVar(value="本部")
        self.ref_point_combo = ttk.Combobox(self, textvariable=self.ref_point_var,
                                          values=list(REFERENCE_POINTS.keys()),
                                          state="readonly", width=10)
        
        # 本日の値入力
        self.today_coords_label = ttk.Label(self, text="本日の値 (経度,緯度,標高):")
        self.today_lng_var = tk.StringVar(value="136.555")
        self.today_lat_var = tk.StringVar(value="36.072")
        self.today_alt_var = tk.StringVar(value="0")
        
        self.today_lng_entry = ttk.Entry(self, textvariable=self.today_lng_var, width=12)
        self.today_lat_entry = ttk.Entry(self, textvariable=self.today_lat_var, width=10)
        self.today_alt_entry = ttk.Entry(self, textvariable=self.today_alt_var, width=8)
        
        # コピーボタン
        self.copy_button = ttk.Button(self, text="コピー", command=self.copy_reference_data, width=8)
        
        # 初期状態設定
        self.update_ctrl()
        self.update_yaw()
        self.update_hover()
        self.update_deviation()
    
    def on_height_change(self, event=None):
        if HEIGHT_OPTIONS.get(self.hc.get()) == "custom":
            self.he.config(state="normal")
            self.he.delete(0, tk.END)
            self.he.focus()
        else:
            self.he.config(state="disabled")
            self.he.delete(0, tk.END)
    
    def update_ctrl(self):
        if self.ph.get():
            self.vd.set(False)
        if self.vd.get():
            self.ph.set(False)
        
        if self.ph.get() or self.vd.get():
            self.gc.state(["disabled"])
            self.gm.set(True)
        else:
            self.gc.state(["!disabled"])
        
        if self.vd.get():
            self.vd_suffix_label.grid(row=2, column=2, sticky="e", padx=(10,2))
            self.vd_suffix_entry.grid(row=2, column=3, sticky="w")
        else:
            self.vd_suffix_label.grid_forget()
            self.vd_suffix_entry.grid_forget()
    
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
        """ホバリング設定の表示/非表示を制御"""
        if self.hv.get():
            self.hover_time_label.grid(row=6, column=1, sticky="e", padx=(10,2))
            self.hover_time_entry.grid(row=6, column=2, sticky="w")
        else:
            self.hover_time_label.grid_forget()
            self.hover_time_entry.grid_forget()
    
    def update_deviation(self):
        """偏差補正設定の表示/非表示を制御"""
        if self.dc.get():
            self.ref_point_label.grid(row=7, column=1, sticky="e", padx=(10,2))
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
        """基準値をクリップボードにコピー"""
        try:
            ref_point = self.ref_point_var.get()
            ref_coords = REFERENCE_POINTS[ref_point]
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            copy_text = f"{current_time},{ref_coords[0]},{ref_coords[1]},{ref_coords[2]}"
            pyperclip.copy(copy_text)
            messagebox.showinfo("コピー完了", f"クリップボードにコピーしました:\n{copy_text}")
        except Exception as e:
            messagebox.showerror("エラー", f"コピーに失敗しました: {e}")
    
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
        if self.yf.get():
            yval = YAW_OPTIONS.get(self.yc.get())
            if yval == "custom":
                try:
                    yaw_angle = float(self.ye.get())
                except:
                    pass
            else:
                yaw_angle = float(yval)
        
        # ホバリング時間の取得
        hover_time = 0
        if self.hv.get():
            try:
                hover_time = max(0, float(self.hover_time_var.get()))
            except:
                hover_time = 2.0  # デフォルト値
        
        # 偏差補正の取得
        coordinate_deviation = None
        if self.dc.get():
            try:
                ref_point = self.ref_point_var.get()
                ref_coords = REFERENCE_POINTS[ref_point]
                current_lng = float(self.today_lng_var.get())
                current_lat = float(self.today_lat_var.get())
                current_alt = float(self.today_alt_var.get())
                
                current_coords = (current_lng, current_lat, current_alt)
                deviation = calculate_deviation(ref_coords, current_coords)
                
                # 安全性チェック
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
            "do_gimbal": self.gm.get(),
            "yaw_fix": self.yf.get(),
            "yaw_angle": yaw_angle,
            "speed": max(1, min(15, self.sp.get())),
            "sensor_modes": [m for m, var in self.sm_vars.items() if var.get()],
            "hover_time": hover_time,
            "coordinate_deviation": coordinate_deviation
        }

def main():
    root = TkinterDnD.Tk()
    root.title("ATL→ASL 変換＋撮影制御ツール (ver. GUI32 - ヨー角制御最適化版)")
    root.geometry("800x750")
    
    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)
    
    app = AppGUI(frm)
    app.pack(fill="x", pady=(0,10))
    
    drop = tk.Label(frm, text=".kmzをここにドロップ", bg="lightgray", width=70, height=5, relief=tk.RIDGE)
    drop.pack(pady=12, fill="x")
    drop.drop_target_register(DND_FILES)
    
    log_frame = ttk.LabelFrame(frm, text="ログ")
    log_frame.pack(fill="both", expand=True)
    
    log = scrolledtext.ScrolledText(log_frame, height=16)
    log.pack(fill="both", expand=True)
    
    def on_drop(event):
        path = event.data.strip("{}")
        if not path.lower().endswith(".kmz"):
            messagebox.showwarning("警告", ".kmzファイルのみ対応しています。")
            return
        
        params = app.get_params()
        if params is None:  # エラーが発生した場合
            return
        
        params.update({"path": path, "log": log})
        threading.Thread(target=process_kmz, kwargs=params, daemon=True).start()
    
    drop.dnd_bind("<<Drop>>", on_drop)
    root.mainloop()

if __name__ == "__main__":
    main()